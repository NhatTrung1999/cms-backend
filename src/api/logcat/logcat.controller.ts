import {
  Body,
  Controller,
  Get,
  Post,
  Query,
  Request,
  Res,
  StreamableFile,
} from '@nestjs/common';
import { LogcatService } from './logcat.service';
import {
  CreateLogCat1And4,
  CreateLogCat5,
  CreateLogCat7,
  CreateLogCat9And12,
} from './dto/create-logcat.dto';
import { getFactortyID, getUserId } from 'src/helper/common.helper';
import { Response } from 'express';

@Controller('logcat')
export class LogcatController {
  constructor(private readonly logcatService: LogcatService) {}

  @Get('get-log-cat1-4')
  async getLogCat1And4(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.logcatService.getLogCat1And4(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Post('create-log-cat1-4')
  async createLogCat1And4(@Body() data: CreateLogCat1And4[], @Request() req) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.logcatService.createLogCat1And4(data, factory, userid);
  }

  @Get('export-excel-cat1-4')
  async exportExcelCat1And4(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Res({ passthrough: true }) res: Response,
  ): Promise<StreamableFile> {
    const buffer = await this.logcatService.exportExcelCat1And4(
      dateFrom,
      dateTo,
      factory,
    );

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="LogCat1And4.xlsx"`,
      'Content-Length': buffer.byteLength,
    });
    return new StreamableFile(Buffer.from(buffer));
  }

  // Logging CAT5
  @Get('get-log-cat5')
  async getLogCat5(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.logcatService.getLogCat5(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Post('create-log-cat5')
  async createLogCat5(@Body() data: CreateLogCat5[], @Request() req) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.logcatService.createLogCat5(data, factory, userid);
  }

  @Get('export-excel-cat5')
  async exportExcelCat5(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Res({ passthrough: true }) res: Response,
  ): Promise<StreamableFile> {
    const buffer = await this.logcatService.exportExcelCat5(
      dateFrom,
      dateTo,
      factory,
    );

    res.set({
      'Content-Type':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="LogCat5.xlsx"`,
      'Content-Length': buffer.byteLength,
    });
    return new StreamableFile(Buffer.from(buffer));
  }

  // Logging CAT7
  @Get('get-log-cat7')
  async getLogCat7(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.logcatService.getLogCat7(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Post('create-log-cat7')
  async createLogCat7(@Body() data: CreateLogCat7[], @Request() req) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.logcatService.createLogCat7(data, factory, userid);
  }

  // Logging CAT9&12
  @Get('get-log-cat9-12')
  async getLogCat9And12(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.logcatService.getLogCat9And12(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Post('create-log-cat9-12')
  async createLogCat9And12(@Body() data: CreateLogCat9And12[], @Request() req) {
    const factory = getFactortyID(req);
    const userid = getUserId(req);
    return this.logcatService.createLogCat9And12(data, factory, userid);
  }
}
