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
import { getADataExcelFactoryCat7 } from 'src/helper/cat7.helper';

@Injectable()
export class FilemanagementService {
  private rootFolder: string;
  constructor(
    private configService: ConfigService,
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
    @Inject('LYF_ERP') private readonly LYF_ERP: Sequelize,
    private readonly eventsGateway: EventsGateway,
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
    @Inject('HRIS') private readonly HRIS: Sequelize,
  ) {
    this.rootFolder = this.configService.get(
      'EXCEL_STORAGE_PATH',
      'D:/FileExcel',
    );
    if (!fs.existsSync(this.rootFolder)) {
      fs.mkdirSync(this.rootFolder, { recursive: true });
    }
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
  ) {
    // console.log(module, dateFrom, dateTo, factory, userID);
    const id = uuidv4();
    const fileName = `${factory}-${module}-${dateFrom}-${dateTo}.xlsx`;
    const folder = path.join(this.rootFolder, module);
    if (!fs.existsSync(folder)) {
      fs.mkdirSync(folder, { recursive: true });
    }
    const filePath = path.join(folder, fileName);

    await this.createFileExcel(id, module, filePath, dateFrom, dateTo, factory);

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
  ) {
    // console.log(id, module, filePath, dateFrom, dateTo, factory);
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(module);

    switch (module.toLowerCase()) {
      case 'cat1andcat4':
        await this.fileExcelCat1AndCat4(sheet, dateFrom, dateTo, factory);
        break;
      case 'cat5':
        await this.fileExcelCat5(sheet, dateFrom, dateTo, factory);
        break;
      case 'cat6':
        await this.fileExcelCat6();
        break;
      case 'cat7':
        await this.fileExcelCat7(sheet, dateFrom, dateTo, factory);
        break;
      case 'cat9andcat12':
        // console.log('cat9andcat12');
        await this.fileExcelCat9AndCat12(sheet, dateFrom, dateTo, factory);
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
      default:
        return await this.getAllExcelFactoryCat9AndCat12Data(
          sheet,
          dateFrom,
          dateTo,
        );
    }

    await getADataExcelFactoryCat9AndCat12(sheet, db, dateFrom, dateTo);
  }

  private async getAllExcelFactoryCat9AndCat12Data(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
  ) {
    let where = 'WHERE 1=1';
    let where1 = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
      where1 += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo, dateFrom, dateTo);
    }
    const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE          AS [Date]
                                    ,im.INV_NO            AS Invoice_Number
                                    ,id.STYLE_NAME        AS Article_Name
                                    ,p.Qty                AS Quantity
                                    ,p.GW                 AS Gross_Weight
                                    ,im.CUSTID            AS Customer_ID
                                    ,'Truck'              AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN ISNULL(LEFT(LTRIM(RTRIM(im.INV_NO)) ,3) ,'')='' THEN ''
                                          WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                          WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)<>'AM-' THEN 'VNCLP'
                                          ELSE 'MMRGN'
                                      END                  AS Port_Of_Departure
                                    ,pc.PortCode          AS Port_Of_Arrival
                                    ,CAST('0' AS INT)     AS Land_Transport_Distance
                                    ,CAST('0' AS INT)     AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)     AS Air_Transport_Distance
                                    ,'SEA'                AS Transport_Method
                                    ,CAST('0' AS INT)     AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)     AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)     AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                      LEFT JOIN INVOICE_D  AS id
                                          ON  id.INV_NO = im.INV_NO
                                      LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS) Qty
                                                    ,SUM(GW) GW
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
                                        SELECT CustomerNumber, PortCode, TransportMethod
                                        FROM CMW.CMW.dbo.CMW_PortCode
                                        WHERE TransportMethod = 'SEA'
                                      ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE       AS [Date]
                                    ,is1.Inv_No        AS Invoice_Number
                                    ,'SAMPLE SHOE'     AS Article_Name
                                    ,is1.Qty           AS Quantity
                                    ,is1.GW            AS Gross_Weight
                                    ,im.CUSTID         AS Customer_ID
                                    ,'Truck'           AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN ISNULL(LEFT(LTRIM(RTRIM(im.INV_NO)) ,3) ,'')='' THEN ''
                                          WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                          WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)<>'AM-' THEN 'VNCLP'
                                          ELSE 'MMRGN'
                                      END               AS Port_Of_Departure
                                    ,pc.PortCode       AS Port_Of_Arrival
                                    ,CAST('0' AS INT)  AS Land_Transport_Distance
                                    ,CAST('0' AS INT)  AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)  AS Air_Transport_Distance
                                    ,'AIR'             AS Transport_Method
                                    ,CAST('0' AS INT)  AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)  AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)  AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                      LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                      LEFT JOIN (
                                        SELECT CustomerNumber, PortCode, TransportMethod
                                        FROM CMW.CMW.dbo.CMW_PortCode
                                        WHERE TransportMethod = 'AIR'
                                      ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where1}
                          ) AS Cat9AndCat12`;
    const connects = [this.LYV_ERP, this.LHG_ERP, this.LYM_ERP, this.LVL_ERP];
    const dataResults = await Promise.all(
      connects.map((conn) => {
        return conn.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });
      }),
    );

    const allData = dataResults.flat();

    let data = allData.map((item, index) => ({ ...item, No: index + 1 }));

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
        header: 'Invoice Number',
        key: 'Invoice_Number',
      },
      {
        header: 'Article Name',
        key: 'Article_Name',
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
    let where = 'WHERE 1=1';
    const replacements: any[] = [];
    if (dateTo && dateFrom) {
      where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }
    const query = `SELECT dwo.WASTE_DATE AS Waste_disposal_date
                        ,dtv.TREATMENT_VENDOR_NAME          AS Vender_Name
                        ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
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
        header: 'Vender Name',
        key: 'Vender_Name',
      },
      {
        header: 'Waste collection address',
        key: 'Waste_collection_address',
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
        );
    }

    await getADataExcelFactoryCat1AndCat4(sheet, db, dateFrom, dateTo);
  }

  private async getAllDataExcelFactoryCat1AndCat4(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
  ) {
    let where = 'WHERE 1=1';
    const replacements: any[] = [];
    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR, c.USERDate, 23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }
    const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
                          ,c.USERDate        AS [Date]
                          ,c.CGNO            AS Purchase_Order
                          ,c2.CLBH           AS Material_No
                          ,CAST('0' AS INT)     AS [Weight]
                          ,z.SupplierCode AS Supplier_Code
                          ,z.ThirdCountryLandTransport AS Thirdcountry_Land_Transport
                          ,z.PortofDeparture AS Port_Of_Departure
                          ,z.PortofArrival AS Port_Of_Arrival
                          ,z.Transportationmethod AS Factory_Domestic_Land_Transport
                          ,CAST('0' AS INT)     AS Land_Transport_Distance
                          ,z.SeaTransportDistance AS Sea_Transport_Distance
                          ,CAST('0' AS INT) AS Air_Transport_Distance
                          ,CAST('0' AS INT) AS Land_Transport_Ton_Kilometers
                          ,CAST('0' AS INT) AS Sea_Transport_Ton_Kilometers
                          ,CAST('0' AS INT) AS Air_Transport_Ton_Kilometers
                    FROM   CGZL              AS c
                          INNER JOIN CGZLS  AS c2
                                ON  c2.CGNO = c.CGNO
                          LEFT JOIN zszl    AS z
                                ON  z.zsdh = c.CGNO
                    ${where}`;
    const connects = [this.LYV_ERP, this.LHG_ERP, this.LYM_ERP, this.LVL_ERP];
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
      { header: 'No.', key: 'No' },
      { header: 'Date', key: 'Date' },
      { header: 'Purchase Order', key: 'Purchase_Order' },
      { header: 'Material No.', key: 'Material_No' },
      { header: 'Weight (Unit: Ton)', key: 'Weight' },
      { header: 'Supplier Code', key: 'Supplier_Code' },
      {
        header: 'Third-country Land Transport (A)',
        key: 'Thirdcountry_Land_Transport',
      },
      { header: 'Port of Departure', key: 'Port_Of_Departure' },
      { header: 'Port of Arrival', key: 'Port_Of_Arrival' },
      {
        header: 'Factory (Domestic Land Transport) (B)',
        key: 'Factory_Domestic_Land_Transport',
      },
      {
        header: 'Land Transport Distance (A+B)',
        key: 'Land_Transport_Distance',
      },
      { header: 'Sea Transport Distance', key: 'Sea_Transport_Distance' },
      { header: 'Air Transport Distance', key: 'Air_Transport_Distance' },
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

  async fileExcelCat6() {}

  async fileExcelCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    let db: Sequelize;
    switch (factory.trim()) {
      case 'LYV':
        db = this.HRIS;
        break;
      case 'LHG':
        return 'LHG Coming soon...';
      // db = this.LHG_ERP;
      // break;
      case 'LYM':
        return 'LYM Coming soon...';
      // db = this.LYM_ERP;
      // break;
      case 'LVL':
        return 'LVL Coming soon...';
      // db = this.LVL_ERP;
      // break;
      default:
        return 'Coming soon...';
    }

    await getADataExcelFactoryCat7(sheet, db, dateFrom, dateTo);
  }
}
