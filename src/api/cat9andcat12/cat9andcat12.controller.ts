import {
  Controller,
  Get,
  Query,
  UseInterceptors,
  BadRequestException,
  UploadedFile,
  Post,
} from '@nestjs/common';
import { Cat9andcat12Service } from './cat9andcat12.service';
import { Public } from 'src/decorators';
import { FileInterceptor } from '@nestjs/platform-express';

@Controller('cat9-and-cat12')
export class Cat9andcat12Controller {
  constructor(private readonly cat9andcat12Service: Cat9andcat12Service) {}

  @Get('get-data-cat9-and-cat12')
  async getTest(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return await this.cat9andcat12Service.getData(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Get('get-port-code')
  async getPortCode(
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat9andcat12Service.getPortCode(sortField, sortOrder);
  }

  @Post('import-excel-port-code')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 50 * 1024 * 1024 } }),
  )
  async importExcelPortCode(@UploadedFile() file: Express.Multer.File) {
    if (!file) throw new BadRequestException('No file uploaded!');

    const data = await this.cat9andcat12Service.importExcelPortCode(file);
    return data;
  }

  @Public()
  @Get('auto-sent-cms')
  async autoSentCMS(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('dockey') dockeyCMS: string,
  ) {
    const data = await this.cat9andcat12Service.autoSentCMS(
      dateFrom,
      dateTo,
      factory,
      dockeyCMS,
    );
    return data;
  }

  @Get('get-logging-cat9-and-cat12')
  async getLoggingCat7(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat9andcat12Service.getLoggingCat9AndCat12(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }
}
