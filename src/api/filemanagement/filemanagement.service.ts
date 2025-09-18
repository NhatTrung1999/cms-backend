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

@Injectable()
export class FilemanagementService {
  private rootFolder: string;
  constructor(
    private configService: ConfigService,
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('ERP') private readonly ERP: Sequelize,
  ) {
    this.rootFolder = this.configService.get(
      'EXCEL_STORAGE_PATH',
      'D:/FileExcel',
    );
    if (!fs.existsSync(this.rootFolder)) {
      fs.mkdirSync(this.rootFolder, { recursive: true });
    }
  }

  async getData(module: string, file_name: string, userID: string) {
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
    const fileName = `${id}.xlsx`;
    const filePath = await this.createFile(
      module,
      fileName,
      this.rootFolder,
      date,
    );
    return await this.EIP.query(
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
          '1',
          userID,
          'LYV',
          new Date().toISOString().slice(0, 10),
        ],
        type: QueryTypes.SELECT,
      },
    );
  }

  async createFile(
    module: string,
    fileName: string,
    rootFolder: string,
    date?: string,
  ) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet 1');

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
        await this.excelCat9AndCat12(sheet, date);
        break;
      default:
        console.log(`Unknown excel file module: ${module}`);
        break;
    }

    const folder = path.join(rootFolder, module);
    if (!fs.existsSync(folder)) {
      fs.mkdirSync(folder, { recursive: true });
    }
    const filePath = path.join(folder, fileName);
    await workbook.xlsx.writeFile(filePath);
    return filePath;
  }

  async excelCat9AndCat12(sheet: ExcelJS.Worksheet, date?: string) {
    const payload: any = await this.ERP
      .query(`SELECT im.INV_DATE,im.INV_NO,id.STYLE_NAME,p.Qty,p.GW,im.CUSTID,
            'Truck' LocalLandTransportation,sb.Place_Delivery,im.TO_WHERE Country
            ,ISNULL(ISNULL(bg.SHPIDS,CASE WHEN (do.ShipMode='Air') AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                          WHEN (do.ShipMode='Air Expres') AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                            WHEN (do.ShipMode='Ocean') AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                            WHEN do.ShipMode_1 IS NULL THEN ''
                                            ELSE do.ShipMode_1 END),y.ShipMode) TransportMethod
              
        FROM INVOICE_M im
        LEFT JOIN Ship_Booking sb ON sb.INV_NO = im.INV_NO
        LEFT JOIN INVOICE_D AS id ON id.INV_NO = im.INV_NO
        LEFT JOIN (SELECT INV_NO,RYNO,SUM(PAIRS)Qty,SUM (GW) GW
                  FROM PACKING
                  GROUP BY INV_NO,RYNO) p ON p.INV_NO=id.INV_NO AND p.RYNO=id.RYNO
        LEFT JOIN YWDD y ON y.DDBH=id.RYNO
        LEFT JOIN DE_ORDERM do ON do.ORDERNO=y.YSBH
        LEFT JOIN B_GradeOrder bg ON bg.ORDER_B=y.YSBH
        WHERE 1=1 ${date ? `AND CONVERT(VARCHAR, im.INV_DATE, 23) = '${date}'` : ''}`);

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

    // console.log(payload[0]);

    const data: IDataCat9AndCat12[] = payload[0]?.map((item, index) => ({
      No: index + 1,
      Date: item.INV_DATE,
      Invoice_Number: item.INV_NO,
      Article_Name: item.STYLE_NAME,
      Quantity: item.Qty,
      Gross_Weight: item.GW,
      Customer_ID: item.CUSTID,
      Local_Land_Transportation: item.LocalLandTransportation,
      Port_Of_Departure: item.Place_Delivery,
      Port_Of_Arrival: item.Country,
      Land_Transport_Distance: '',
      Sea_Transport_Distance: '',
      Air_Transport_Distance: '',
      Transport_Method: item.TransportMethod,
      Land_Transport_Ton_Kilometers: '',
      Sea_Transport_Ton_Kilometers: '',
      Air_Transport_Ton_Kilometers: '',
    }));

    data.forEach((item) => sheet.addRow(item));

    sheet.columns.forEach((column) => {
      let maxLength = 10;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? cell.value.toString() : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength + 2;
    });

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
