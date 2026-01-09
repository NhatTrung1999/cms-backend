import { Body, Controller, Get, Post, Query, Request } from '@nestjs/common';
import { LogcatService } from './logcat.service';
import {
  CreateLogCat5,
  CreateLogCat7,
  CreateLogCat9And12,
} from './dto/create-logcat.dto';
import { getFactortyID, getUserId } from 'src/helper/common.helper';

@Controller('logcat')
export class LogcatController {
  constructor(private readonly logcatService: LogcatService) {}

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
