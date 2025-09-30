import {
  Controller,
  Get,
  Request,
  Query,
  UseInterceptors,
  BadRequestException,
  UploadedFile,
} from '@nestjs/common';
import { Cat9andcat12Service } from './cat9andcat12.service';
import { getUserId } from 'src/helper/common.helper';
import { Public } from 'src/decorators';
import { FileInterceptor } from '@nestjs/platform-express';
import * as multer from 'multer';
import * as ExcelJS from 'exceljs';

@Controller('cat9-and-cat12')
export class Cat9andcat12Controller {
  constructor(private readonly cat9andcat12Service: Cat9andcat12Service) {}

  @Get('get-data-cat9-and-cat12')
  async getTest(
    @Query('date') date: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return await this.cat9andcat12Service.getData(
      date,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Get('get-port-code')
  async getPortCode() {
    return this.cat9andcat12Service.getPortCode();
  }

  @Get('import-excel-port-code')
  @UseInterceptors(
    FileInterceptor('file', {
      storage: multer.memoryStorage(),
      fileFilter: (req, file, cb) => {
        if (!file.originalname.match(/\.(xlsx|xls)$/)) {
          return cb(new BadRequestException('Only allow file Excel'), false);
        }
        cb(null, true);
      },
    }),
  )
  async importExcelPortCode(@UploadedFile() file: Express.Multer.File) {
    if (!file) throw new BadRequestException('No file uploaded!');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.buffer.buffer as any);

    const worksheet = workbook.worksheets[0];
    worksheet.eachRow((row, rowNumber) => {
      console.log(row, rowNumber);
    });
  }
}
