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
import { IDataCat9AndCat12 } from 'src/types/cat9andcat12';
import { EventsGateway } from '../events/events.gateway';

@Injectable()
export class FilemanagementService {
  private rootFolder: string;
  constructor(
    private configService: ConfigService,
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('ERP') private readonly ERP: Sequelize,
    private readonly eventsGateway: EventsGateway,
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
    sortField: string = 'Module',
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
          FROM CMS_File_Management
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
        FROM CMS_File_Management
        WHERE ID = ?`,
      {
        replacements: [id],
        type: QueryTypes.SELECT,
      },
    );
    if (payload.length === 0) return false;
    return payload[0];
  }

  async generateFileExcel(module: string, date: string, userID: string) {
    const id = uuidv4();
    const fileName = `${module}-${id}.xlsx`;
    const folder = path.join(this.rootFolder, module);
    if (!fs.existsSync(folder)) {
      fs.mkdirSync(folder, { recursive: true });
    }
    const filePath = path.join(folder, fileName);

    let statusWaiting = await this.fileExcelWaiting(
      id,
      module,
      fileName,
      filePath,
      userID,
    );
    await this.createFileExcel(id, module, filePath, date);
    if (statusWaiting.length === 0) false;
    return true;
  }

  async fileExcelWaiting(
    id: string,
    module: string,
    fileName: string,
    filePath: string,
    userID: string,
  ) {
    try {
      await this.EIP.query(
        `INSERT INTO CMS_File_Management
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
            'LYV',
            new Date().toISOString().slice(0, 10),
          ],
          type: QueryTypes.INSERT,
        },
      );
      const result = await this.EIP.query(
        `SELECT * FROM CMS_File_Management WHERE ID = ?`,
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
        `UPDATE CMS_File_Management
          SET
            [Status] = '1'
          WHERE ID = ?`,
        { replacements: [id], type: QueryTypes.UPDATE },
      );
      const result = await this.EIP.query(
        `SELECT * FROM CMS_File_Management WHERE ID = ?`,
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
    date?: string,
  ) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(module);

    switch (module.toLowerCase()) {
      case 'cat1andcat4':
        console.log('cat1andcat4');
        break;
      case 'cat5':
        console.log('cat5');
        break;
      case 'cat6':
        console.log('cat6');
        break;
      case 'cat7':
        console.log('cat7');
        break;
      case 'cat9andcat12':
        await this.fileExcelCat9AndCat12(sheet, date);
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
              () => {
                return 'Export success!';
              },
            );
          }
        }, 5000);
      })
      .catch((error: any) => {
        return 'Error export!';
      });
  }

  async fileExcelCat9AndCat12(sheet: ExcelJS.Worksheet, date?: string) {
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
                    ${where}`;
    const data = await this.ERP.query(query, {
      replacements,
      type: QueryTypes.SELECT,
    });
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
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
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
