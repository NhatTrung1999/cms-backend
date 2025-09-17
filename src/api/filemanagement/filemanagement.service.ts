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

@Injectable()
export class FilemanagementService {
  private rootFolder: string;
  constructor(
    private configService: ConfigService,
    @Inject('EIP') private readonly EIP: Sequelize,
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
        where += ` AND Module = '${module}'`;
      }

      if (file_name !== '') {
        where += ` AND File_Name = '${file_name}'`;
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
    // console.log(payload);
    return payload[0];
  }

  async generateFileExcel(module: string, date: string, userID: string) {
    const id = uuidv4();
    const fileName = `${id}.xlsx`;
    const filePath = await this.createFile(module, fileName, this.rootFolder);
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
          date,
        ],
        type: QueryTypes.SELECT,
      },
    );
  }

  async createFile(module: string, fileName: string, rootFolder: string) {
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
        await this.excelCat9AndCat12(sheet);
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

  async excelCat9AndCat12(sheet: ExcelJS.Worksheet) {
    sheet.columns = [
      { header: 'Tháng', key: 'month', width: 15 },
      { header: 'Doanh thu', key: 'revenue', width: 20 },
      { header: 'Lợi nhuận', key: 'profit', width: 20 },
    ];

    const data: any[] = [
      { month: 'Tháng 7', revenue: 10000000, profit: 3000000 },
      { month: 'Tháng 8', revenue: 12000000, profit: 4000000 },
    ];

    data.forEach((item) => sheet.addRow(item));
    sheet.getRow(1).eachCell((cell) => {
      ((cell.font = { bold: true, color: { argb: 'ffffff' } }),
        (cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF007ACC' },
        }));
    });
  }
}
