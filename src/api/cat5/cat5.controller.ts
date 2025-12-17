import { Controller, Get, Query } from '@nestjs/common';
import { Cat5Service } from './cat5.service';
import { Public } from 'src/decorators';

@Controller('cat5')
export class Cat5Controller {
  constructor(private readonly cat5Service: Cat5Service) {}

  @Get('get-data-cat5')
  async getDataWMS(
    @Query('dateFrom') dateFrom: string,
    @Query('dateTo') dateTo: string,
    @Query('factory') factory: string,
    @Query('page') page: number,
    @Query('limit') limit: number,
    @Query('sortField') sortField: string,
    @Query('sortOrder') sortOrder: string,
  ) {
    // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
    return this.cat5Service.getDataWMS(
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
    ) {
      const data = await this.cat5Service.autoSentCMS(dateFrom, dateTo);
      return data;
    }
}
