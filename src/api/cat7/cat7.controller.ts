import { Controller, Get, Query } from '@nestjs/common';
import { Cat7Service } from './cat7.service';
import { Public } from 'src/decorators';

@Controller('cat7')
export class Cat7Controller {
  constructor(private readonly cat7Service: Cat7Service) {}

  @Get('get-data-cat7')
  async getDataCat7(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat7Service.getDataCat7(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Get('custom-export')
  async getCustomExport(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat7Service.getCustomExport(
      dateFrom,
      dateTo,
      factory,
      +page,
      +limit,
      sortField,
      sortOrder,
    );
  }

  @Public()
  @Get('auto-sent-cms')
  async autoSentCMS(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
  ) {
    const data = await this.cat7Service.autoSentCMS(dateFrom, dateTo, factory);
    return data;
  }

  @Get('get-logging-cat7')
  async getLoggingCat7(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    return this.cat7Service.getLoggingCat7(
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
